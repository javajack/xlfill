package xlfill

import "fmt"

// AutoColWidthCommand implements the jx:autoColWidth command.
// It processes its inner area, then estimates column widths based on
// content length and sets them.
type AutoColWidthCommand struct {
	Area *Area
}

func (c *AutoColWidthCommand) Name() string { return "autoColWidth" }
func (c *AutoColWidthCommand) Reset()       {}

// newAutoColWidthCommandFromAttrs creates an AutoColWidthCommand from parsed attributes.
func newAutoColWidthCommandFromAttrs(attrs map[string]string) (Command, error) {
	return &AutoColWidthCommand{}, nil
}

// ApplyAt processes the inner area, then sets column widths based on content.
func (c *AutoColWidthCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, nil
	}

	size, err := c.Area.ApplyAt(cellRef, ctx)
	if err != nil {
		return ZeroSize, err
	}

	etx := unwrapExcelizeTransformer(transformer)
	if etx == nil {
		return size, nil
	}

	// For each column, scan output cells and compute max width
	for col := 0; col < size.Width; col++ {
		maxLen := 8.0 // minimum width
		for row := 0; row < size.Height; row++ {
			ref := NewCellRef(cellRef.Sheet, cellRef.Row+row, cellRef.Col+col)
			val, _ := etx.file.GetCellValue(cellRef.Sheet, ref.CellName())
			width := float64(len(val)) * 1.2
			if width > maxLen {
				maxLen = width
			}
		}
		if maxLen > 60 {
			maxLen = 60 // cap
		}
		colName := ColToName(cellRef.Col + col)
		if err := etx.file.SetColWidth(cellRef.Sheet, colName, colName, maxLen); err != nil {
			return size, fmt.Errorf("set col width %s: %w", colName, err)
		}
	}

	return size, nil
}
