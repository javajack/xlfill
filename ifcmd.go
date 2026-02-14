package xlfill

import "fmt"

// IfCommand implements the jx:if command for conditional rendering.
type IfCommand struct {
	Condition string // boolean expression to evaluate
	IfArea    *Area  // area to render when condition is true
	ElseArea  *Area  // area to render when condition is false (optional)
}

func (c *IfCommand) Name() string { return "if" }
func (c *IfCommand) Reset()       {}

// newIfCommandFromAttrs creates an IfCommand from parsed attributes.
func newIfCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &IfCommand{
		Condition: attrs["condition"],
	}
	if cmd.Condition == "" {
		return nil, fmt.Errorf("if command requires 'condition' attribute")
	}
	return cmd, nil
}

// ApplyAt evaluates the condition and applies the appropriate area.
func (c *IfCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	result, err := ctx.IsConditionTrue(c.Condition)
	if err != nil {
		return ZeroSize, fmt.Errorf("evaluate condition %q: %w", c.Condition, err)
	}

	if result {
		if c.IfArea != nil {
			return c.IfArea.ApplyAt(cellRef, ctx)
		}
	} else {
		if c.ElseArea != nil {
			return c.ElseArea.ApplyAt(cellRef, ctx)
		}
	}

	return ZeroSize, nil
}
