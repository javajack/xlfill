package xlfill

// AutoRowHeightCommand implements jx:autoRowHeight to auto-fit row heights after content is written.
type AutoRowHeightCommand struct {
	Area *Area
}

func (c *AutoRowHeightCommand) Name() string { return "autoRowHeight" }
func (c *AutoRowHeightCommand) Reset()       {}

func newAutoRowHeightCommandFromAttrs(attrs map[string]string) (Command, error) {
	return &AutoRowHeightCommand{}, nil
}

// ApplyAt processes the area and then sets each row to auto-height.
func (c *AutoRowHeightCommand) ApplyAt(cellRef CellRef, ctx *Context, tx Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, nil
	}

	size, err := c.Area.ApplyAt(cellRef, ctx)
	if err != nil {
		return ZeroSize, err
	}

	// Set each output row to auto-height by setting height to -1
	for row := 0; row < size.Height; row++ {
		tx.SetRowHeight(cellRef.Sheet, cellRef.Row+row, -1)
	}

	return size, nil
}
