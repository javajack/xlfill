---
title: API Reference
description: Complete API reference for the XLFill Go library — functions, options, and types.
---

This page covers every public function, option, and type in XLFill. For guided walkthroughs, see the [Getting Started](/guides/getting-started/) guide.

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

## Filler (advanced)

For repeated fills or fine-grained control, create a `Filler`:

```go
filler := xlfill.NewFiller(
    xlfill.WithTemplate("template.xlsx"),
    xlfill.WithClearTemplateCells(true),
    xlfill.WithRecalculateOnOpen(true),
)

err := filler.Fill(data, "output.xlsx")
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
| `WithCommand(name, factory)` | Register a [custom command](/guides/custom-commands/) |
| `WithAreaListener(listener)` | Add a [cell transform hook](/guides/area-listeners/) |
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

Implement this to create [custom commands](/guides/custom-commands/):

```go
type Command interface {
    Name() string
    ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error)
    Reset()
}
```

### AreaListener interface

Implement this for [cell transform hooks](/guides/area-listeners/):

```go
type AreaListener interface {
    BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool
    AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer)
}
```

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

Curious about how XLFill performs at scale?

**[Performance &rarr;](/reference/performance/)**
