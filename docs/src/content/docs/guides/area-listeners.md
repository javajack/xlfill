---
title: Area Listeners
description: Hook into cell processing for conditional styling, logging, and validation.
---

Area listeners let you intercept every cell transformation — before and after — without writing a custom command. Use them for conditional row styling, audit logging, data validation, or any cross-cutting concern.

## The AreaListener interface

```go
type AreaListener interface {
    BeforeTransformCell(src, target CellRef, ctx *Context, tx Transformer) bool
    AfterTransformCell(src, target CellRef, ctx *Context, tx Transformer)
}
```

- **`BeforeTransformCell`** — called before each cell is processed. Return `false` to skip the default transformation (you handle it yourself). Return `true` to proceed normally.
- **`AfterTransformCell`** — called after each cell is processed. Use it for post-processing like styling.

Both receive the source cell (template position), target cell (output position), the data context, and the transformer.

## Registering a listener

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithAreaListener(&MyListener{}),
)
```

## Example: Alternate row colors

```go
type AlternateRowListener struct{}

func (l *AlternateRowListener) BeforeTransformCell(
    src, target xlfill.CellRef,
    ctx *xlfill.Context,
    tx xlfill.Transformer,
) bool {
    return true // proceed with default transformation
}

func (l *AlternateRowListener) AfterTransformCell(
    src, target xlfill.CellRef,
    ctx *xlfill.Context,
    tx xlfill.Transformer,
) {
    if target.Row%2 == 0 {
        // Use transformer API to apply a light background
    }
}
```

## Example: Audit logging

```go
type AuditListener struct{}

func (l *AuditListener) BeforeTransformCell(
    src, target xlfill.CellRef, ctx *xlfill.Context, tx xlfill.Transformer,
) bool {
    log.Printf("Processing cell %s -> %s", src, target)
    return true
}

func (l *AuditListener) AfterTransformCell(
    src, target xlfill.CellRef, ctx *xlfill.Context, tx xlfill.Transformer,
) {}
```

## StyleListener: conditional styling

For conditional formatting based on cell values, implement `StyleListener`. It's called after each cell is transformed and lets you override style properties:

```go
type StyleListener interface {
    StyleCell(target CellRef, value any, ctx *Context) *StyleOverride
}

type StyleOverride struct {
    Bold      *bool
    Italic    *bool
    FontColor *string  // hex color e.g. "#FF0000"
    FillColor *string  // hex background color
    FontSize  *float64
}
```

### Example: Red text for negative values

```go
type NegativeHighlighter struct{}

func (l *NegativeHighlighter) BeforeTransformCell(
    src, target xlfill.CellRef, ctx *xlfill.Context, tx xlfill.Transformer,
) bool { return true }

func (l *NegativeHighlighter) AfterTransformCell(
    src, target xlfill.CellRef, ctx *xlfill.Context, tx xlfill.Transformer,
) {}

func (l *NegativeHighlighter) StyleCell(
    target xlfill.CellRef, value any, ctx *xlfill.Context,
) *xlfill.StyleOverride {
    if f, ok := value.(float64); ok && f < 0 {
        color := "FF0000"
        bold := true
        return &xlfill.StyleOverride{FontColor: &color, Bold: &bold}
    }
    return nil // no change
}
```

Register it the same way as any listener:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithAreaListener(&NegativeHighlighter{}),
)
```

A listener can implement both `AreaListener` and `StyleListener` — the two are independent call paths.

**Thread safety note:** When used with `WithParallelism`, listeners are called from multiple goroutines. Use atomic operations or mutexes for any shared mutable state (counters, accumulators, etc.).

## PreWrite callback

For logic that runs **after all template processing** but **before writing the output**, use `WithPreWrite`:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithPreWrite(func(tx xlfill.Transformer) error {
        // Set print area, add final calculations, etc.
        return nil
    }),
)
```

## Template sheet control

Related options for controlling the output:

```go
// Keep the template sheet in output (default: removed)
xlfill.WithKeepTemplateSheet(true)

// Hide the template sheet instead of removing
xlfill.WithHideTemplateSheet(true)

// Don't clear unexpanded ${...} expressions (default: cleared)
xlfill.WithClearTemplateCells(false)
```

## What's next?

Area listeners are one of several debugging tools. For the complete toolkit — including `Validate()`, `Describe()`, and common troubleshooting tips:

**[Debugging & Troubleshooting &rarr;](/xlfill/guides/debugging/)**

For the complete list of functions, options, and types:

**[API Reference &rarr;](/xlfill/reference/api/)**
