---
title: API Reference
description: Complete API reference for the XLFill Go library — functions, options, and types.
---

This page covers every public function, option, and type in the XLFill library. For guided walkthroughs, see the [Getting Started](/xlfill/guides/getting-started/) guide.

## Top-level functions

These are the simplest way to use XLFill. One function call, done.

### Fill

```go
func Fill(templatePath, outputPath string, data map[string]any, opts ...Option) error
```

Read a template file, fill it with data, write the output file.

```go
xlfill.Fill("template.xlsx", "report.xlsx", data)
```

### FillBytes

```go
func FillBytes(templatePath string, data map[string]any, opts ...Option) ([]byte, error)
```

Read a template file, fill it, return the result as bytes. Useful when you need the output in memory.

```go
bytes, err := xlfill.FillBytes("template.xlsx", data)
```

### FillReader

```go
func FillReader(template io.Reader, output io.Writer, data map[string]any, opts ...Option) error
```

Fill from an `io.Reader`, write to an `io.Writer`. Perfect for HTTP handlers — no temp files needed:

```go
func handler(w http.ResponseWriter, r *http.Request) {
    tmpl, _ := os.Open("template.xlsx")
    defer tmpl.Close()
    xlfill.FillReader(tmpl, w, data)
}
```

### Validate

```go
func Validate(templatePath string, opts ...Option) ([]ValidationIssue, error)
```

Check a template for structural and expression errors **without requiring data**. Returns a list of issues found. Use this in CI pipelines or during development to catch problems early.

```go
issues, err := xlfill.Validate("template.xlsx")
if err != nil {
    log.Fatal(err) // template couldn't be opened or parsed at all
}
for _, issue := range issues {
    fmt.Println(issue) // [ERROR] Sheet1!B2: invalid expression syntax "e.Name +": ...
}
```

What it checks:
- **Expression syntax** — validates all `${...}` in cell values and formulas
- **Command attributes** — validates `items`, `condition`, `select`, `headers`, `data` expressions
- **Bounds** — verifies each command's `lastCell` fits within its parent area

### Describe

```go
func Describe(templatePath string, opts ...Option) (string, error)
```

Parse a template and return a human-readable tree showing the area hierarchy, commands with attributes, and expressions found in cells. Useful for understanding what the engine "sees" when it reads your template.

```go
output, err := xlfill.Describe("template.xlsx")
fmt.Print(output)
```

Sample output:
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

## Filler (advanced)

For repeated fills or fine-grained control, create a `Filler`:

```go
filler := xlfill.NewFiller(
    xlfill.WithTemplate("template.xlsx"),
    xlfill.WithClearTemplateCells(true),
    xlfill.WithRecalculateOnOpen(true),
)

err := filler.Fill(data, "output.xlsx")

// Validate and Describe also available on Filler
issues, err := filler.Validate()
description, err := filler.Describe()
```

## Options

All options work with both the top-level functions and `NewFiller`.

### Template source

| Option | Description |
|--------|-------------|
| `WithTemplate(path)` | Set template file path |
| `WithTemplateReader(r io.Reader)` | Set template from a reader |

### Expression configuration

| Option | Description |
|--------|-------------|
| `WithExpressionNotation(begin, end)` | Custom delimiters (default: `${`, `}`) |

### Output control

| Option | Description |
|--------|-------------|
| `WithClearTemplateCells(bool)` | Clear unexpanded `${...}` cells (default: `true`) |
| `WithKeepTemplateSheet(bool)` | Keep template sheet in output (default: `false`) |
| `WithHideTemplateSheet(bool)` | Hide template sheet instead of removing (default: `false`) |
| `WithRecalculateOnOpen(bool)` | Tell Excel to recalculate formulas on open |

### Extensibility

| Option | Description |
|--------|-------------|
| `WithCommand(name, factory)` | Register a [custom command](/xlfill/guides/custom-commands/) |
| `WithAreaListener(listener)` | Add a [cell transform hook](/xlfill/guides/area-listeners/) |
| `WithPreWrite(fn)` | Callback before writing output |

## Types

### CellRef

A reference to a single cell:

```go
type CellRef struct {
    SheetName string
    Col       int  // 0-based column index
    Row       int  // 1-based row number
}
```

### AreaRef

A rectangular range of cells:

```go
type AreaRef struct {
    SheetName    string
    FirstCellRef CellRef
    LastCellRef  CellRef
}
```

### Size

The dimensions of a processed area:

```go
type Size struct {
    Width  int
    Height int
}
```

### Command interface

Implement this to create [custom commands](/xlfill/guides/custom-commands/):

```go
type Command interface {
    Name() string
    ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error)
    Reset()
}
```

### AreaListener interface

Implement this for [cell transform hooks](/xlfill/guides/area-listeners/):

```go
type AreaListener interface {
    BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool
    AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer)
}
```

### ValidationIssue

Returned by `Validate()`:

```go
type Severity int
const (
    SeverityError   Severity = iota // template will fail at runtime
    SeverityWarning                 // template may produce unexpected results
)

type ValidationIssue struct {
    Severity Severity
    CellRef  CellRef
    Message  string
}
```

`String()` formats as `[ERROR] Sheet1!A2: message` or `[WARN] Sheet1!A2: message`.

## Data input

The `data` parameter accepts `map[string]any`. Values can be:

| Type | Example | Template access |
|------|---------|----------------|
| Primitives | `string`, `int`, `float64`, `bool` | `${name}`, `${count}` |
| Maps | `map[string]any` | `${employee.Name}` |
| Slices | `[]any`, `[]Employee` | Used in `jx:each(items="...")` |
| Structs | Any Go struct | `${emp.Name}` — fields by name |
| Byte slices | `[]byte` | Used in `jx:image(src="...")` |

Nested access works via dot notation: `${employee.Address.City}`.

## What's next?

Having trouble with a template? See the debugging toolkit:

**[Debugging & Troubleshooting &rarr;](/xlfill/guides/debugging/)**

Curious about how XLFill performs at scale?

**[Performance &rarr;](/xlfill/reference/performance/)**
