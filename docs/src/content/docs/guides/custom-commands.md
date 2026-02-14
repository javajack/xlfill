---
title: Custom Commands
description: Extend XLFill with your own template commands.
---

XLFill ships with a rich set of built-in commands, but every project is different. Custom commands let you add your own `jx:` directives to handle domain-specific logic.

## The Command interface

Every command — built-in or custom — implements this interface:

```go
type Command interface {
    Name() string
    ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error)
    Reset()
}
```

- **`Name()`** — the command name used after `jx:` in templates
- **`ApplyAt()`** — called when XLFill processes this command. Receives the cell position, the data context, and the transformer for reading/writing cells. Returns the `Size` of the output area.
- **`Reset()`** — called before each fill operation to reset state

## Registering a custom command

Use `WithCommand` with a factory function. The factory receives the attributes from the template comment and returns a command instance:

```go
filler := xlfill.NewFiller(
    xlfill.WithTemplate("template.xlsx"),
    xlfill.WithCommand("highlight", func(attrs map[string]string) (xlfill.Command, error) {
        return &HighlightCommand{
            Color: attrs["color"],
        }, nil
    }),
)
```

Now you can use it in templates:

```
jx:highlight(color="yellow" lastCell="C1")
```

## Full example

A command that sets a background color on a cell range:

```go
type HighlightCommand struct {
    Color string
}

func (c *HighlightCommand) Name() string { return "highlight" }

func (c *HighlightCommand) ApplyAt(
    cellRef xlfill.CellRef,
    ctx *xlfill.Context,
    tx xlfill.Transformer,
) (xlfill.Size, error) {
    // Use the transformer to read cells, apply styles, write values
    // Return the size of the area produced
    return xlfill.Size{Width: 3, Height: 1}, nil
}

func (c *HighlightCommand) Reset() {}
```

## Parsing attributes

The factory function receives all template attributes as `map[string]string`. Parse them as needed:

```go
xlfill.WithCommand("repeat", func(attrs map[string]string) (xlfill.Command, error) {
    count, err := strconv.Atoi(attrs["count"])
    if err != nil {
        return nil, fmt.Errorf("invalid count: %w", err)
    }
    return &RepeatCommand{Count: count}, nil
})
```

Template usage: `jx:repeat(count="3" lastCell="C1")`

## What's next?

For hooking into cell processing without creating a full command:

**[Area Listeners &rarr;](/xlfill/guides/area-listeners/)**
