---
title: Debugging & Troubleshooting
description: Tools and techniques for inspecting templates, validating expressions, and diagnosing issues in XLFill.
---

XLFill provides built-in tools for catching template issues early and understanding what the engine sees when it processes your template. This page covers every debugging technique available.

## Validate: catch errors without data

`Validate()` checks your template for structural and expression errors **without requiring any data**. Run it in CI, in tests, or during development to catch problems before they reach production.

```go
issues, err := xlfill.Validate("template.xlsx")
if err != nil {
    // Template couldn't be opened or parsed at all
    log.Fatal(err)
}
for _, issue := range issues {
    fmt.Println(issue)
}
```

Output:
```
[ERROR] Sheet1!B2: invalid expression syntax "e.Name +": unexpected token "+"
[ERROR] Sheet1!A3: each command has invalid items expression "employees[": unexpected token "["
```

### What Validate checks

| Check | What it catches |
|-------|----------------|
| Expression syntax | Bad `${...}` in cell values — e.g., `${e.Name +}` |
| Formula expressions | Bad `${...}` inside formulas — e.g., `=SUM(${bad syntax})` |
| Command attributes | Invalid expressions in `items`, `condition`, `select`, `headers`, `data` |
| Bounds | A command's `lastCell` extends beyond its parent `jx:area` |
| Structural errors | Missing `jx:area`, invalid cell references (returned as `error`, not issues) |

### Use in CI / tests

```go
func TestTemplateValid(t *testing.T) {
    issues, err := xlfill.Validate("templates/monthly_report.xlsx")
    require.NoError(t, err)
    assert.Empty(t, issues, "template has validation issues: %v", issues)
}
```

### ValidationIssue type

Each issue includes a severity, cell reference, and message:

```go
type ValidationIssue struct {
    Severity Severity  // SeverityError or SeverityWarning
    CellRef  CellRef
    Message  string
}
```

`issue.String()` formats as `[ERROR] Sheet1!A2: message` or `[WARN] Sheet1!A2: message`.

## Describe: see what the engine sees

When a template doesn't produce the output you expect, `Describe()` shows you exactly what the engine parsed — the area hierarchy, command attributes, and expressions found in each cell.

```go
output, err := xlfill.Describe("template.xlsx")
if err != nil {
    log.Fatal(err)
}
fmt.Print(output)
```

Output:
```
Template: template.xlsx
Sheet1!A1:C2 area (3x2)
  Commands:
    Sheet1!A2 each (3x1) items="employees" var="e"
      Sheet1!A2:C2 area (3x1)
        Expressions:
          A2: ${e.Name}
          B2: ${e.Age}
          C2: ${e.Salary}
```

### What to look for

- **Missing commands** — if a `jx:each` doesn't appear, the comment text may be malformed
- **Wrong area bounds** — if the area dimensions look off, check your `lastCell` attribute
- **Missing expressions** — if `${...}` cells don't appear, they may be outside the area bounds
- **Unexpected nesting** — commands should nest inside their parent area, not siblings

### Nested template example

For templates with nested loops or conditionals, `Describe` shows the full tree:

```
Template: report.xlsx
Sheet1!A1:C5 area (3x5)
  Expressions:
    A1: ${title}
  Commands:
    Sheet1!A2 each (3x4) items="departments" var="dept"
      Sheet1!A2:C5 area (3x4)
        Expressions:
          A2: ${dept.Name}
        Commands:
          Sheet1!A3 each (3x1) items="dept.Employees" var="e"
            Sheet1!A3:C3 area (3x1)
              Expressions:
                A3: ${e.Name}
                B3: ${e.Role}
                C3: ${e.Salary}
          Sheet1!C5 if (1x1) condition="dept.ShowTotal"
```

## Error messages: reading the error chain

When `Fill()` fails at runtime, the error message includes the full context chain. Here's how to read it:

```
process area at Sheet1!A1: command each (template Sheet1!A2) at target Sheet1!A5:
  select filter "e.Active" at item 3: expression evaluation failed: ...
```

Breaking this down:

| Part | Meaning |
|------|---------|
| `process area at Sheet1!A1` | The root area being processed |
| `command each (template Sheet1!A2)` | The command that failed, and which template cell it came from |
| `at target Sheet1!A5` | The output cell where the command was being applied |
| `select filter "e.Active" at item 3` | The specific operation and iteration index |

The **template cell** tells you where to look in your `.xlsx` file. The **target cell** tells you where in the output the failure occurred. The **item index** tells you which data record triggered the error.

## AreaListener: trace every cell transformation

For deep debugging, implement `AreaListener` to log every cell as it's processed:

```go
type DebugListener struct{}

func (l *DebugListener) BeforeTransformCell(
    src, target xlfill.CellRef,
    ctx *xlfill.Context,
    tx xlfill.Transformer,
) bool {
    cd := tx.GetCellData(src)
    if cd != nil && cd.Value != nil {
        log.Printf("CELL %s -> %s  value=%v", src, target, cd.Value)
    }
    return true // proceed with default transformation
}

func (l *DebugListener) AfterTransformCell(
    src, target xlfill.CellRef,
    ctx *xlfill.Context,
    tx xlfill.Transformer,
) {}
```

Register it:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithAreaListener(&DebugListener{}),
)
```

This logs every cell copy from source to target, showing the template expression or value. Useful for understanding the processing order and spotting which cell causes an issue.

## PreWrite callback: inspect final state

`WithPreWrite` runs after all template processing but before writing the output. Use it to inspect or modify the final transformer state:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithPreWrite(func(tx xlfill.Transformer) error {
        // Inspect a specific cell in the output
        cd := tx.GetCellData(xlfill.NewCellRef("Sheet1", 0, 0))
        log.Printf("A1 final value: %v", cd.Value)
        return nil
    }),
)
```

## Common issues and fixes

### "no jx:area found"

Your template has no root `jx:area` command. Every template needs at least one cell comment containing `jx:area(lastCell="...")`.

**Fix:** Add a `jx:area` comment to the top-left cell of your template region.

### Expressions not being replaced

The `${...}` cells show up as literal text in the output.

**Possible causes:**
1. The cell is outside the `jx:area` bounds — run `Describe()` to check the area dimensions
2. Custom notation was set but the template uses default `${...}` — check your `WithExpressionNotation` option
3. `WithClearTemplateCells(false)` is set — unreplaced expressions won't be cleared

### Command not executing

A `jx:each` or `jx:if` exists in a comment but nothing happens.

**Possible causes:**
1. The comment is a threaded comment, not a note — XLFill reads cell **notes**, not threaded comments (see the [Getting Started](/xlfill/guides/getting-started/) guide for how to add notes in each editor)
2. The command syntax has a typo — run `Validate()` to check
3. The command is outside any `jx:area` — commands must be inside an area's bounds

### "lastCell" out of bounds

The error says a command's area extends beyond its parent.

**Fix:** Make sure the `lastCell` attribute in the child command doesn't exceed the `lastCell` of the parent `jx:area`. Run `Validate()` to catch this at build time.

### Wrong loop output order

Items appear in unexpected order.

**Possible causes:**
1. Go maps don't guarantee order — if your data source is a map, the iteration order is random
2. Use `orderBy` to sort: `jx:each(items="employees" var="e" orderBy="e.Name ASC" lastCell="...")`

### Formula not expanding

A formula like `=SUM(A2:A2)` doesn't expand to cover all generated rows.

**Fix:** The formula must reference cells **within** the `jx:each` area. The formula cell itself must be **outside** the loop but **inside** the `jx:area`. See the [Formulas guide](/xlfill/guides/formulas/) for details.

## Debugging checklist

When something isn't working, go through this in order:

1. **`Validate()`** — catches syntax errors, bad expressions, and bounds issues without needing data
2. **`Describe()`** — shows the parsed template structure; verify it matches your intent
3. **Check the error message** — read the full chain: area, command, template cell, target cell, iteration index
4. **`AreaListener`** — trace cell-by-cell processing to find exactly where things go wrong
5. **`PreWrite`** — inspect the final output state before it's written to file
6. **Open the template** — sometimes the simplest fix is to open the `.xlsx` and check that comments are on the right cells

## What's next?

For the complete list of functions, options, and types:

**[API Reference &rarr;](/xlfill/reference/api/)**
