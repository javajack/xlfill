package xlfill

import (
	"fmt"
	"strconv"
)

// PageBreakCommand implements the jx:pageBreak command.
// It processes its inner area, then inserts a horizontal page break
// after the last output row.
type PageBreakCommand struct {
	Area *Area
}

func (c *PageBreakCommand) Name() string { return "pageBreak" }
func (c *PageBreakCommand) Reset()       {}

// newPageBreakCommandFromAttrs creates a PageBreakCommand from parsed attributes.
func newPageBreakCommandFromAttrs(attrs map[string]string) (Command, error) {
	return &PageBreakCommand{}, nil
}

// ApplyAt processes the inner area, then inserts a page break after the last output row.
func (c *PageBreakCommand) ApplyAt(cellRef CellRef, ctx *Context, transformer Transformer) (Size, error) {
	if c.Area == nil {
		return ZeroSize, nil
	}

	size, err := c.Area.ApplyAt(cellRef, ctx)
	if err != nil {
		return ZeroSize, err
	}

	// Insert page break after last output row.
	// The break cell is one row below the last output row.
	breakCell := ColToName(cellRef.Col) + strconv.Itoa(cellRef.Row+size.Height+1)
	etx := unwrapExcelizeTransformer(transformer)
	if etx != nil {
		if err := etx.file.InsertPageBreak(cellRef.Sheet, breakCell); err != nil {
			return size, fmt.Errorf("insert page break: %w", err)
		}
	}

	return size, nil
}
