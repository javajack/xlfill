package xlfill

import "fmt"

// GroupCommand implements the jx:group command.
// It processes its inner area normally, then registers a deferred action
// that groups the expanded output rows in Excel (outline grouping).
type GroupCommand struct {
	Collapsed string // "true" or "false" (default "false")
	Area      *Area
}

func (c *GroupCommand) Name() string { return "group" }
func (c *GroupCommand) Reset()       {}

// newGroupCommandFromAttrs creates a GroupCommand from parsed attributes.
func newGroupCommandFromAttrs(attrs map[string]string) (Command, error) {
	cmd := &GroupCommand{
		Collapsed: attrs["collapsed"],
	}
	if cmd.Collapsed == "" {
		cmd.Collapsed = "false"
	}
	return cmd, nil
}

// ApplyAt processes the inner area, then registers a deferred action for row grouping.
func (c *GroupCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, nil
	}

	// Process the inner area normally first
	size, err := c.Area.ApplyAt(cellRef, ctx)
	if err != nil {
		return ZeroSize, err
	}

	if size.Height <= 0 {
		return size, nil
	}

	// Capture values for the closure
	sheet := cellRef.Sheet
	startRow := cellRef.Row
	endRow := cellRef.Row + size.Height - 1
	collapsed := c.Collapsed == "true"

	// Register deferred action
	ctx.RegisterDeferred(DeferredAction{
		Name:     "group",
		Sheet:    sheet,
		StartRow: startRow,
		StartCol: cellRef.Col,
		EndRow:   endRow,
		EndCol:   cellRef.Col + size.Width - 1,
		Execute: func(tx *ExcelizeTransformer) error {
			// excelize SetRowOutlineLevel uses 1-based row numbers
			for row := startRow; row <= endRow; row++ {
				if err := tx.file.SetRowOutlineLevel(sheet, row+1, 1); err != nil {
					return fmt.Errorf("set outline level for row %d: %w", row+1, err)
				}
			}

			// If collapsed, hide the grouped rows
			if collapsed {
				for row := startRow; row <= endRow; row++ {
					if err := tx.file.SetRowVisible(sheet, row+1, false); err != nil {
						return fmt.Errorf("hide row %d: %w", row+1, err)
					}
				}
			}

			return nil
		},
	})

	return size, nil
}
