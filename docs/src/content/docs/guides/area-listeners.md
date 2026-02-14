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

For the complete list of functions, options, and types:

**[API Reference &rarr;](/reference/api/)**
