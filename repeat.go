package xlfill

import "fmt"

// MaxRepeatCount limits the maximum number of repeat iterations to prevent resource exhaustion.
const MaxRepeatCount = 1_000_000

// RepeatCommand implements the jx:repeat command for repeating an area a fixed number of times.
// Unlike jx:each which requires a collection, jx:repeat simply repeats the area N times,
// optionally exposing the iteration index via the var attribute.
type RepeatCommand struct {
	Count     string // expression evaluating to an integer count
	Var       string // optional: loop index variable name
	Direction string // "DOWN" (default) or "RIGHT"
	Area      *Area
}

func (c *RepeatCommand) Name() string { return "repeat" }
func (c *RepeatCommand) Reset()       {}

// newRepeatCommandFromAttrs creates a RepeatCommand from parsed attributes.
func newRepeatCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &RepeatCommand{
		Count:     attrs["count"],
		Var:       attrs["var"],
		Direction: attrs["direction"],
	}
	if cmd.Count == "" {
		return nil, fmt.Errorf("repeat command requires 'count' attribute")
	}
	if cmd.Direction == "" {
		cmd.Direction = "DOWN"
	}
	if cmd.Direction != "DOWN" && cmd.Direction != "RIGHT" {
		return nil, fmt.Errorf("repeat command: invalid direction %q (must be DOWN or RIGHT)", cmd.Direction)
	}
	return cmd, nil
}

// ApplyAt repeats the area the specified number of times.
func (c *RepeatCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, fmt.Errorf("repeat command has no area")
	}

	countVal, err := ctx.Evaluate(c.Count)
	if err != nil {
		return ZeroSize, fmt.Errorf("evaluate count %q: %w", c.Count, err)
	}

	count := toInt(countVal)
	if count <= 0 {
		return ZeroSize, nil
	}
	if count > MaxRepeatCount {
		return ZeroSize, fmt.Errorf("repeat count %d exceeds maximum %d", count, MaxRepeatCount)
	}

	isRight := c.Direction == "RIGHT"
	totalSize := ZeroSize

	for i := 0; i < count; i++ {
		// Set index variable if configured
		var rv *RunVar
		if c.Var != "" {
			rv = NewRunVar(ctx, c.Var)
			rv.Set(i)
		}

		var iterTarget CellRef
		if isRight {
			iterTarget = NewCellRef(cellRef.Sheet, cellRef.Row, cellRef.Col+totalSize.Width)
		} else {
			iterTarget = NewCellRef(cellRef.Sheet, cellRef.Row+totalSize.Height, cellRef.Col)
		}

		iterSize, err := c.Area.ApplyAt(iterTarget, ctx)
		if rv != nil {
			rv.Close()
		}
		if err != nil {
			return ZeroSize, fmt.Errorf("repeat iteration %d: %w", i, err)
		}

		if isRight {
			totalSize.Width += iterSize.Width
			if iterSize.Height > totalSize.Height {
				totalSize.Height = iterSize.Height
			}
		} else {
			totalSize.Height += iterSize.Height
			if iterSize.Width > totalSize.Width {
				totalSize.Width = iterSize.Width
			}
		}
	}

	return totalSize, nil
}
